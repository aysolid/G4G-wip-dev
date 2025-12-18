/**
 * G4G Research Operations Web App
 * Main Server-Side Code
 *
 * This file contains the core web app functionality including:
 * - Web app entry points (doGet)
 * - Page routing
 * - Session management
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  APP_NAME: 'G4G Research Operations',
  VERSION: '1.0.0',
  SITES: ['UGA', 'Missouri'],
  PERIODS: ['Spring', 'Summer', 'Fall'],
  ROLES: ['admin', 'facilitator', 'viewer'],
  PARTICIPANT_STATUSES: ['active', 'withdrawn', 'completed'],
  CHECKLIST_STATUSES: ['not_started', 'completed', 'missing'],

  // The 18 non-negotiable protocol items
  INSTRUMENTS: [
    { number: 1, name: 'Consent Form', category: 'enrollment' },
    { number: 2, name: 'Assent Form / Pre-test', category: 'enrollment' },
    { number: 3, name: 'Lesson 1 Journal', category: 'lesson' },
    { number: 4, name: 'Lesson 2 Journal', category: 'lesson' },
    { number: 5, name: 'Lesson 3 Journal', category: 'lesson' },
    { number: 6, name: 'Lesson 3 - Design a Game Worksheet', category: 'worksheet' },
    { number: 7, name: 'Lesson 4 Journal', category: 'lesson' },
    { number: 8, name: 'Lesson 4 - Paper Prototyping Worksheet', category: 'worksheet' },
    { number: 9, name: 'Lesson 5 Journal', category: 'lesson' },
    { number: 10, name: 'Lesson 5 - Debugging Worksheet', category: 'worksheet' },
    { number: 11, name: 'Lesson 6 Journal', category: 'lesson' },
    { number: 12, name: 'Lesson 6 - Game Refinement Worksheet', category: 'worksheet' },
    { number: 13, name: 'Lesson 7 Journal', category: 'lesson' },
    { number: 14, name: 'Lesson 7 - Playtesting Feedback Guide', category: 'worksheet' },
    { number: 15, name: 'Lesson 8 Journal', category: 'lesson' },
    { number: 16, name: 'Post-Test', category: 'post' },
    { number: 17, name: 'Participant Feedback Survey', category: 'post' },
    { number: 18, name: 'Parent Satisfaction Survey', category: 'post' }
  ]
};

// ============================================
// WEB APP ENTRY POINTS
// ============================================

/**
 * Main entry point for the web app
 * Handles GET requests and routes to appropriate pages
 */
function doGet(e) {
  const page = e.parameter.page || 'index';
  const sessionToken = e.parameter.token || '';

  // Validate session for protected pages
  const protectedPages = ['dashboard', 'users', 'rollouts', 'participants', 'participant-detail', 'reports'];

  if (protectedPages.includes(page)) {
    const session = validateSession(sessionToken);
    if (!session) {
      return HtmlService.createTemplateFromFile('login')
        .evaluate()
        .setTitle(CONFIG.APP_NAME + ' - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  try {
    const template = HtmlService.createTemplateFromFile(page);
    return template.evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('Error loading page: ' + error.message);
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Include HTML files (for modular HTML)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get the web app URL
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Get configuration for client-side use
 */
function getConfig() {
  return {
    appName: CONFIG.APP_NAME,
    version: CONFIG.VERSION,
    sites: CONFIG.SITES,
    periods: CONFIG.PERIODS,
    roles: CONFIG.ROLES,
    instruments: CONFIG.INSTRUMENTS,
    participantStatuses: CONFIG.PARTICIPANT_STATUSES,
    checklistStatuses: CONFIG.CHECKLIST_STATUSES
  };
}

// ============================================
// DATABASE INITIALIZATION
// ============================================

/**
 * Initialize all database sheets
 * Run this once when setting up the application
 */
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create Users sheet
  createSheetIfNotExists(ss, 'Users', [
    'userId', 'username', 'passwordHash', 'fullName', 'role', 'site', 'status', 'createdAt', 'createdBy', 'lastLogin'
  ]);

  // Create StudyRollouts sheet
  createSheetIfNotExists(ss, 'StudyRollouts', [
    'rolloutId', 'site', 'schoolName', 'period', 'year', 'status', 'createdAt', 'createdBy', 'description'
  ]);

  // Create Participants sheet
  createSheetIfNotExists(ss, 'Participants', [
    'participantId', 'fullName', 'site', 'rolloutId', 'schoolName', 'period', 'year',
    'enrollmentDate', 'enrolledBy', 'status', 'notes', 'completionPercentage'
  ]);

  // Create Checklist sheet
  createSheetIfNotExists(ss, 'Checklist', [
    'checklistId', 'participantId', 'instrumentNumber', 'instrumentName', 'category',
    'status', 'completedDate', 'completedBy', 'notes'
  ]);

  // Create ActivityLog sheet
  createSheetIfNotExists(ss, 'ActivityLog', [
    'logId', 'timestamp', 'userId', 'userName', 'action', 'targetType', 'targetId', 'details'
  ]);

  // Create Sessions sheet (for authentication)
  createSheetIfNotExists(ss, 'Sessions', [
    'sessionId', 'userId', 'token', 'createdAt', 'expiresAt', 'isActive'
  ]);

  // Create Config sheet
  createSheetIfNotExists(ss, 'Config', [
    'key', 'value', 'description', 'updatedAt'
  ]);

  Logger.log('Database initialized successfully!');
  return { success: true, message: 'Database initialized successfully!' };
}

/**
 * Create a sheet if it doesn't exist
 */
function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('Created sheet: ' + sheetName);
  }

  return sheet;
}

/**
 * Create the initial admin user
 * Run this once after initializing the database
 */
function createInitialAdmin() {
  const username = 'admin';
  const password = generateRandomPassword();
  const passwordHash = hashPassword(password);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) {
    throw new Error('Users sheet not found. Run initializeDatabase() first.');
  }

  // Check if admin already exists
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      Logger.log('Admin user already exists!');
      return { success: false, message: 'Admin user already exists' };
    }
  }

  const userId = generateUUID();
  const timestamp = new Date().toISOString();

  usersSheet.appendRow([
    userId,
    username,
    passwordHash,
    'System Administrator',
    'admin',
    'All',
    'active',
    timestamp,
    'system',
    ''
  ]);

  Logger.log('Initial admin created!');
  Logger.log('Username: ' + username);
  Logger.log('Temporary Password: ' + password);
  Logger.log('IMPORTANT: Change this password after first login!');

  return {
    success: true,
    message: 'Admin created successfully',
    username: username,
    password: password
  };
}

/**
 * UTILITY: Generate a password hash
 * Run this function from the script editor to generate hashed passwords
 * for manual entry in the spreadsheet
 *
 * Usage: Change the password below and run this function.
 * Copy the hash from the execution log to your spreadsheet.
 */
function generatePasswordHash() {
  const password = 'admin'; // <-- Change this to your desired password
  const hash = hashPassword(password);
  Logger.log('Password: ' + password);
  Logger.log('Hash: ' + hash);
  Logger.log('Copy the hash above to the passwordHash column in the Users sheet');
  return hash;
}

// ============================================
// AUTHENTICATION FUNCTIONS
// ============================================

/**
 * Authenticate user with username and password
 */
function authenticateUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) {
    return { success: false, message: 'System not initialized' };
  }

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  // Debug logging
  Logger.log('Attempting login for username: ' + username);
  Logger.log('Headers found: ' + JSON.stringify(headers));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUsername = row[headers.indexOf('username')];
    const rowPasswordHash = row[headers.indexOf('passwordHash')];
    const rowStatus = row[headers.indexOf('status')];

    Logger.log('Checking row ' + i + ': username=' + rowUsername + ', status=' + rowStatus);

    if (rowUsername === username && rowStatus === 'active') {
      Logger.log('Username matched! Verifying password...');
      Logger.log('Stored hash: ' + rowPasswordHash);
      Logger.log('Input password hash: ' + hashPassword(password));

      if (verifyPassword(password, rowPasswordHash)) {
        // Create session
        const userId = row[headers.indexOf('userId')];
        const session = createSession(userId);

        // Update last login
        const lastLoginCol = headers.indexOf('lastLogin') + 1;
        usersSheet.getRange(i + 1, lastLoginCol).setValue(new Date().toISOString());

        // Log activity
        logActivity(userId, row[headers.indexOf('fullName')], 'LOGIN', 'user', userId, 'User logged in');

        return {
          success: true,
          token: session.token,
          user: {
            userId: userId,
            username: rowUsername,
            fullName: row[headers.indexOf('fullName')],
            role: row[headers.indexOf('role')],
            site: row[headers.indexOf('site')]
          }
        };
      } else {
        Logger.log('Password verification failed');
      }
    }
  }

  return { success: false, message: 'Invalid username or password' };
}

/**
 * Create a new session for a user
 */
function createSession(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('Sessions');

  const sessionId = generateUUID();
  const token = generateSessionToken();
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + 24 * 60 * 60 * 1000); // 24 hours

  sessionsSheet.appendRow([
    sessionId,
    userId,
    token,
    createdAt.toISOString(),
    expiresAt.toISOString(),
    true
  ]);

  return { sessionId, token, expiresAt };
}

/**
 * Validate a session token
 */
function validateSession(token) {
  if (!token) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('Sessions');

  if (!sessionsSheet) return null;

  const data = sessionsSheet.getDataRange().getValues();
  const headers = data[0];
  const tokenCol = headers.indexOf('token');
  const expiresCol = headers.indexOf('expiresAt');
  const isActiveCol = headers.indexOf('isActive');
  const userIdCol = headers.indexOf('userId');

  for (let i = 1; i < data.length; i++) {
    if (data[i][tokenCol] === token && data[i][isActiveCol] === true) {
      const expiresAt = new Date(data[i][expiresCol]);
      if (expiresAt > new Date()) {
        // Session is valid, get user info
        const userId = data[i][userIdCol];
        return getUserById(userId);
      }
    }
  }

  return null;
}

/**
 * Get current user from session token
 */
function getCurrentUser(token) {
  return validateSession(token);
}

/**
 * Logout user (invalidate session)
 */
function logoutUser(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('Sessions');

  if (!sessionsSheet) return { success: false };

  const data = sessionsSheet.getDataRange().getValues();
  const headers = data[0];
  const tokenCol = headers.indexOf('token');
  const isActiveCol = headers.indexOf('isActive');

  for (let i = 1; i < data.length; i++) {
    if (data[i][tokenCol] === token) {
      sessionsSheet.getRange(i + 1, isActiveCol + 1).setValue(false);
      return { success: true };
    }
  }

  return { success: false };
}

/**
 * Change user password
 */
function changePassword(token, currentPassword, newPassword) {
  const user = validateSession(token);
  if (!user) {
    return { success: false, message: 'Invalid session' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === user.userId) {
      const storedPassword = data[i][headers.indexOf('passwordHash')];

      if (!verifyPassword(currentPassword, storedPassword)) {
        return { success: false, message: 'Current password is incorrect' };
      }

      // Store plain text password since the Google Sheet is highly protected
      usersSheet.getRange(i + 1, headers.indexOf('passwordHash') + 1).setValue(newPassword);

      logActivity(user.userId, user.fullName, 'PASSWORD_CHANGE', 'user', user.userId, 'Password changed');

      return { success: true, message: 'Password changed successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

// ============================================
// USER MANAGEMENT FUNCTIONS
// ============================================

/**
 * Get user by ID
 */
function getUserById(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) return null;

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      return {
        userId: data[i][headers.indexOf('userId')],
        username: data[i][headers.indexOf('username')],
        fullName: data[i][headers.indexOf('fullName')],
        role: data[i][headers.indexOf('role')],
        site: data[i][headers.indexOf('site')],
        status: data[i][headers.indexOf('status')],
        createdAt: data[i][headers.indexOf('createdAt')],
        lastLogin: data[i][headers.indexOf('lastLogin')]
      };
    }
  }

  return null;
}

/**
 * Get all users (Admin only)
 */
function getAllUsers(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) return { success: false, message: 'Users sheet not found' };

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      userId: data[i][headers.indexOf('userId')],
      username: data[i][headers.indexOf('username')],
      fullName: data[i][headers.indexOf('fullName')],
      role: data[i][headers.indexOf('role')],
      site: data[i][headers.indexOf('site')],
      status: data[i][headers.indexOf('status')],
      createdAt: data[i][headers.indexOf('createdAt')],
      lastLogin: data[i][headers.indexOf('lastLogin')]
    });
  }

  return { success: true, users: users };
}

/**
 * Create a new user (Admin only)
 */
function createUser(token, userData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  // Basic validation
  if (!userData.role || !CONFIG.ROLES.includes(userData.role)) {
    return { success: false, message: 'Invalid role selected' };
  }

  if (!userData.site) {
    return { success: false, message: 'Site is required' };
  }

  const validSites = CONFIG.SITES.concat(['All']);
  if (validSites.indexOf(userData.site) === -1) {
    return { success: false, message: 'Invalid site selected' };
  }

  if (userData.role === 'facilitator' && userData.site === 'All') {
    return { success: false, message: 'Facilitators must be assigned to a single site' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  // Check if username already exists
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('username')] === userData.username) {
      return { success: false, message: 'Username already exists' };
    }
  }

  const userId = generateUUID();
  // Use provided password or generate one if not provided
  const password = userData.password || generateRandomPassword();
  const timestamp = new Date().toISOString();

  usersSheet.appendRow([
    userId,
    userData.username,
    password,  // Plain text password
    userData.fullName,
    userData.role,
    userData.site,
    'active',
    timestamp,
    currentUser.userId,
    ''
  ]);

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_USER', 'user', userId,
    'Created user: ' + userData.username);

  return {
    success: true,
    message: 'User created successfully',
    tempPassword: password,
    userId: userId
  };
}

/**
 * Update user (Admin only)
 */
function updateUser(token, userId, userData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      const existingRole = data[i][headers.indexOf('role')];
      const nextRole = userData.role || existingRole;

      if (userData.fullName) {
        usersSheet.getRange(i + 1, headers.indexOf('fullName') + 1).setValue(userData.fullName);
      }
      if (userData.role) {
        if (!CONFIG.ROLES.includes(userData.role)) {
          return { success: false, message: 'Invalid role selected' };
        }
        usersSheet.getRange(i + 1, headers.indexOf('role') + 1).setValue(userData.role);
      }
      if (userData.site) {
        const validSites = CONFIG.SITES.concat(['All']);
        if (validSites.indexOf(userData.site) === -1) {
          return { success: false, message: 'Invalid site selected' };
        }
        if (nextRole === 'facilitator' && userData.site === 'All') {
          return { success: false, message: 'Facilitators must be assigned to a single site' };
        }
        usersSheet.getRange(i + 1, headers.indexOf('site') + 1).setValue(userData.site);
      }
      if (userData.status) {
        usersSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(userData.status);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_USER', 'user', userId,
        'Updated user: ' + data[i][headers.indexOf('username')]);

      return { success: true, message: 'User updated successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

/**
 * Reset user password (Admin only)
 */
function resetUserPassword(token, userId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      const newPassword = generateRandomPassword();
      // Store plain text password since the Google Sheet is highly protected
      usersSheet.getRange(i + 1, headers.indexOf('passwordHash') + 1).setValue(newPassword);

      logActivity(currentUser.userId, currentUser.fullName, 'RESET_PASSWORD', 'user', userId,
        'Reset password for: ' + data[i][headers.indexOf('username')]);

      return { success: true, tempPassword: newPassword };
    }
  }

  return { success: false, message: 'User not found' };
}

/**
 * Delete a user account (Admin only)
 */
function deleteUser(token, userId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  if (currentUser.userId === userId) {
    return { success: false, message: 'You cannot delete your own account' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const sessionsSheet = ss.getSheetByName('Sessions');

  if (!usersSheet) {
    return { success: false, message: 'Users sheet not found' };
  }

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const userIdCol = headers.indexOf('userId');
  const usernameCol = headers.indexOf('username');
  const roleCol = headers.indexOf('role');

  // Ensure at least one admin remains
  let adminCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][roleCol] === 'admin') {
      adminCount++;
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdCol] === userId) {
      const username = data[i][usernameCol];
      const role = data[i][roleCol];

      if (role === 'admin' && adminCount <= 1) {
        return { success: false, message: 'At least one admin user is required' };
      }

      usersSheet.deleteRow(i + 1);

      // Deactivate all sessions for this user
      if (sessionsSheet) {
        const sessionData = sessionsSheet.getDataRange().getValues();
        const sessionHeaders = sessionData[0];
        const sUserIdCol = sessionHeaders.indexOf('userId');
        const sIsActiveCol = sessionHeaders.indexOf('isActive');

        for (let j = 1; j < sessionData.length; j++) {
          if (sessionData[j][sUserIdCol] === userId) {
            sessionsSheet.getRange(j + 1, sIsActiveCol + 1).setValue(false);
          }
        }
      }

      logActivity(currentUser.userId, currentUser.fullName, 'DELETE_USER', 'user', userId,
        'Deleted user: ' + username);

      return { success: true, message: 'User deleted successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

// ============================================
// STUDY ROLLOUT FUNCTIONS
// ============================================

/**
 * Get all study rollouts
 */
function getAllRollouts(token, siteFilter) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return { success: false, message: 'StudyRollouts sheet not found' };

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const rollouts = [];

  for (let i = 1; i < data.length; i++) {
    const rollout = {
      rolloutId: data[i][headers.indexOf('rolloutId')],
      site: data[i][headers.indexOf('site')],
      schoolName: data[i][headers.indexOf('schoolName')],
      period: data[i][headers.indexOf('period')],
      year: data[i][headers.indexOf('year')],
      status: data[i][headers.indexOf('status')],
      createdAt: data[i][headers.indexOf('createdAt')],
      description: data[i][headers.indexOf('description')]
    };

    // Apply site filter if specified
    if (!siteFilter || siteFilter === 'All' || rollout.site === siteFilter) {
      rollouts.push(rollout);
    }
  }

  return { success: true, rollouts: rollouts };
}

/**
 * Get rollouts by site
 */
function getRolloutsBySite(token, site) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && currentUser.site !== site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  return getAllRollouts(token, site);
}

/**
 * Create a new study rollout (Admin only)
 */
function createRollout(token, rolloutData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized - Admin access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  const rolloutId = generateUUID();
  const timestamp = new Date().toISOString();

  rolloutsSheet.appendRow([
    rolloutId,
    rolloutData.site,
    rolloutData.schoolName,
    rolloutData.period,
    rolloutData.year,
    'active',
    timestamp,
    currentUser.userId,
    rolloutData.description || ''
  ]);

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_ROLLOUT', 'rollout', rolloutId,
    'Created rollout: ' + rolloutData.schoolName + ' (' + rolloutData.period + ' ' + rolloutData.year + ')');

  return { success: true, message: 'Rollout created successfully', rolloutId: rolloutId };
}

/**
 * Update study rollout (Admin only)
 */
function updateRollout(token, rolloutId, rolloutData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('rolloutId')] === rolloutId) {
      if (rolloutData.schoolName) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('schoolName') + 1).setValue(rolloutData.schoolName);
      }
      if (rolloutData.period) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('period') + 1).setValue(rolloutData.period);
      }
      if (rolloutData.year) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('year') + 1).setValue(rolloutData.year);
      }
      if (rolloutData.status) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(rolloutData.status);
      }
      if (rolloutData.description !== undefined) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('description') + 1).setValue(rolloutData.description);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_ROLLOUT', 'rollout', rolloutId,
        'Updated rollout');

      return { success: true, message: 'Rollout updated successfully' };
    }
  }

  return { success: false, message: 'Rollout not found' };
}

/**
 * Get rollout by ID
 */
function getRolloutById(rolloutId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return null;

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('rolloutId')] === rolloutId) {
      return {
        rolloutId: data[i][headers.indexOf('rolloutId')],
        site: data[i][headers.indexOf('site')],
        schoolName: data[i][headers.indexOf('schoolName')],
        period: data[i][headers.indexOf('period')],
        year: data[i][headers.indexOf('year')],
        status: data[i][headers.indexOf('status')],
        description: data[i][headers.indexOf('description')]
      };
    }
  }

  return null;
}

// ============================================
// PARTICIPANT FUNCTIONS
// ============================================

/**
 * Get all participants with optional filters
 */
function getAllParticipants(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');

  if (!participantsSheet) return { success: false, message: 'Participants sheet not found' };

  const data = participantsSheet.getDataRange().getValues();
  const headers = data[0];
  const participants = [];

  filters = filters || {};

  for (let i = 1; i < data.length; i++) {
    const participant = {
      participantId: data[i][headers.indexOf('participantId')],
      fullName: data[i][headers.indexOf('fullName')],
      site: data[i][headers.indexOf('site')],
      rolloutId: data[i][headers.indexOf('rolloutId')],
      schoolName: data[i][headers.indexOf('schoolName')],
      period: data[i][headers.indexOf('period')],
      year: data[i][headers.indexOf('year')],
      enrollmentDate: data[i][headers.indexOf('enrollmentDate')],
      enrolledBy: data[i][headers.indexOf('enrolledBy')],
      status: data[i][headers.indexOf('status')],
      notes: data[i][headers.indexOf('notes')],
      completionPercentage: data[i][headers.indexOf('completionPercentage')] || 0
    };

    // Apply filters
    let include = true;

    if (filters.site && filters.site !== 'All' && participant.site !== filters.site) {
      include = false;
    }
    if (filters.rolloutId && filters.rolloutId !== 'All' && participant.rolloutId !== filters.rolloutId) {
      include = false;
    }
    if (filters.status && participant.status !== filters.status) {
      include = false;
    }
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      const matchesSearch =
        participant.fullName.toLowerCase().includes(searchLower) ||
        participant.participantId.toLowerCase().includes(searchLower) ||
        participant.schoolName.toLowerCase().includes(searchLower);
      if (!matchesSearch) include = false;
    }

    // Facilitators see their site by default unless viewing all
    if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && participant.site !== currentUser.site) {
      include = false;
    }

    if (include) {
      participants.push(participant);
    }
  }

  return { success: true, participants: participants };
}

/**
 * Get participant by ID with checklist
 */
function getParticipantById(token, participantId) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Get participant data
  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  let participant = null;

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('participantId')] === participantId) {
      participant = {
        participantId: pData[i][pHeaders.indexOf('participantId')],
        fullName: pData[i][pHeaders.indexOf('fullName')],
        site: pData[i][pHeaders.indexOf('site')],
        rolloutId: pData[i][pHeaders.indexOf('rolloutId')],
        schoolName: pData[i][pHeaders.indexOf('schoolName')],
        period: pData[i][pHeaders.indexOf('period')],
        year: pData[i][pHeaders.indexOf('year')],
        enrollmentDate: pData[i][pHeaders.indexOf('enrollmentDate')],
        enrolledBy: pData[i][pHeaders.indexOf('enrolledBy')],
        status: pData[i][pHeaders.indexOf('status')],
        notes: pData[i][pHeaders.indexOf('notes')],
        completionPercentage: pData[i][pHeaders.indexOf('completionPercentage')] || 0
      };
      break;
    }
  }

  if (!participant) {
    return { success: false, message: 'Participant not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && participant.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  // Get checklist items
  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];
  const checklist = [];

  for (let i = 1; i < cData.length; i++) {
    if (cData[i][cHeaders.indexOf('participantId')] === participantId) {
      checklist.push({
        checklistId: cData[i][cHeaders.indexOf('checklistId')],
        participantId: cData[i][cHeaders.indexOf('participantId')],
        instrumentNumber: cData[i][cHeaders.indexOf('instrumentNumber')],
        instrumentName: cData[i][cHeaders.indexOf('instrumentName')],
        category: cData[i][cHeaders.indexOf('category')],
        status: cData[i][cHeaders.indexOf('status')],
        completedDate: cData[i][cHeaders.indexOf('completedDate')],
        completedBy: cData[i][cHeaders.indexOf('completedBy')],
        notes: cData[i][cHeaders.indexOf('notes')]
      });
    }
  }

  // Sort checklist by instrument number
  checklist.sort((a, b) => a.instrumentNumber - b.instrumentNumber);

  return { success: true, participant: participant, checklist: checklist };
}

/**
 * Enroll a new participant (Admin and Facilitator)
 */
function enrollParticipant(token, participantData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  // Get rollout info
  const rollout = getRolloutById(participantData.rolloutId);
  if (!rollout) {
    return { success: false, message: 'Invalid rollout selected' };
  }

  if (currentUser.role === 'facilitator') {
    if (!currentUser.site || currentUser.site === 'All') {
      return { success: false, message: 'Facilitators must be assigned to a single site before enrolling participants' };
    }

    if (rollout.site !== currentUser.site) {
      return { success: false, message: 'Unauthorized to enroll participants for this site' };
    }
  }

  // Generate participant ID
  const participantId = generateParticipantId(rollout.site);
  const timestamp = new Date().toISOString();

  // Add participant
  participantsSheet.appendRow([
    participantId,
    participantData.fullName,
    rollout.site,
    participantData.rolloutId,
    rollout.schoolName,
    rollout.period,
    rollout.year,
    timestamp,
    currentUser.userId,
    'active',
    participantData.notes || '',
    0 // completionPercentage
  ]);

  // Create checklist items for all 18 instruments
  CONFIG.INSTRUMENTS.forEach(instrument => {
    const checklistId = generateUUID();
    checklistSheet.appendRow([
      checklistId,
      participantId,
      instrument.number,
      instrument.name,
      instrument.category,
      'not_started',
      '',
      '',
      ''
    ]);
  });

  logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT', 'participant', participantId,
    'Enrolled: ' + participantData.fullName);

  return {
    success: true,
    message: 'Participant enrolled successfully',
    participantId: participantId
  };
}

/**
 * Update participant info
 */
function updateParticipant(token, participantId, participantData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const canChangeStatus = currentUser.role === 'admin';

  if (participantData.status && !canChangeStatus) {
    return { success: false, message: 'Only administrators can change participant status' };
  }

  if (participantData.status && CONFIG.PARTICIPANT_STATUSES.indexOf(participantData.status) === -1) {
    return { success: false, message: 'Invalid status selected' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const data = participantsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('participantId')] === participantId) {
      if (currentUser.role === 'facilitator' && currentUser.site !== 'All' &&
          data[i][headers.indexOf('site')] !== currentUser.site) {
        return { success: false, message: 'Unauthorized for this site' };
      }

      if (participantData.fullName) {
        participantsSheet.getRange(i + 1, headers.indexOf('fullName') + 1).setValue(participantData.fullName);
      }
      if (participantData.status) {
        participantsSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(participantData.status);
      }
      if (participantData.notes !== undefined) {
        participantsSheet.getRange(i + 1, headers.indexOf('notes') + 1).setValue(participantData.notes);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_PARTICIPANT', 'participant', participantId,
        'Updated participant info');

      return { success: true, message: 'Participant updated successfully' };
    }
  }

  return { success: false, message: 'Participant not found' };
}

/**
 * Delete a participant and their checklist entries (Admin only)
 */
function deleteParticipant(token, participantId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet) {
    return { success: false, message: 'Participants sheet not found' };
  }

  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const participantIdCol = pHeaders.indexOf('participantId');
  let participantName = '';
  let participantSite = '';
  let rowIndex = -1;

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][participantIdCol] === participantId) {
      participantName = pData[i][pHeaders.indexOf('fullName')];
      participantSite = pData[i][pHeaders.indexOf('site')];
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: 'Participant not found' };
  }

  participantsSheet.deleteRow(rowIndex);

  if (checklistSheet) {
    const cData = checklistSheet.getDataRange().getValues();
    const cHeaders = cData[0];
    const checklistParticipantCol = cHeaders.indexOf('participantId');

    for (let i = cData.length - 1; i >= 1; i--) {
      if (cData[i][checklistParticipantCol] === participantId) {
        checklistSheet.deleteRow(i + 1);
      }
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'DELETE_PARTICIPANT', 'participant', participantId,
    'Deleted participant: ' + participantName + ' (' + participantSite + ')');

  return { success: true, message: 'Participant deleted successfully' };
}

/**
 * Generate participant ID with site prefix
 */
function generateParticipantId(site) {
  const prefix = site === 'UGA' ? 'UGA' : 'MIZ';
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + '-' + timestamp + random;
}

/**
 * Generate a CSV template for batch participant enrollment
 */
function generateEnrollmentTemplate(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  if (!rolloutsSheet) {
    return { success: false, message: 'Study rollouts sheet not found' };
  }

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const rollouts = [];
  const isFacilitator = currentUser.role === 'facilitator' && currentUser.site && currentUser.site !== 'All';

  for (let i = 1; i < data.length; i++) {
    const site = data[i][headers.indexOf('site')];
    if (isFacilitator && site !== currentUser.site) continue;
    rollouts.push({
      rolloutId: data[i][headers.indexOf('rolloutId')],
      site: site,
      schoolName: data[i][headers.indexOf('schoolName')],
      period: data[i][headers.indexOf('period')],
      year: data[i][headers.indexOf('year')],
      status: data[i][headers.indexOf('status')]
    });
  }

  const csvHeaders = ['fullName', 'rolloutName'];
  const rows = [csvHeaders];

  rollouts.forEach(r => {
    const label = `${r.schoolName} (${r.period} ${r.year})`;
    rows.push([
      '',
      label
    ]);
  });

  if (rows.length === 1) {
    rows.push(['', 'Example Rollout (Period Year)']);
  }

  const csv = rows.map(row => row.map(csvEscape).join(',')).join('\n');

  return {
    success: true,
    csv: csv,
    filename: 'participant_enrollment_template.csv',
    rolloutCount: rollouts.length
  };
}

/**
 * Import participants from CSV (batch enrollment/upsert)
 */
function importParticipantsCSV(token, fileData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  if (!fileData || !fileData.csv) {
    return { success: false, message: 'No CSV content provided' };
  }

  let rows;
  try {
    rows = Utilities.parseCsv(fileData.csv);
  } catch (error) {
    return { success: false, message: 'Unable to parse CSV: ' + error.message };
  }

  if (!rows || rows.length === 0) {
    return { success: false, message: 'CSV is empty' };
  }

  const headers = rows[0].map(h => h.trim());
  const required = ['fullName', 'rolloutName'];
  const missing = required.filter(col => headers.indexOf(col) === -1);
  if (missing.length > 0) {
    return { success: false, message: 'Missing required columns: ' + missing.join(', ') };
  }

  const idx = name => headers.indexOf(name);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!participantsSheet || !checklistSheet || !rolloutsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Build rollout map for quick lookups
  const rolloutData = rolloutsSheet.getDataRange().getValues();
  const rolloutHeaders = rolloutData[0];
  const rolloutMap = {};
  const isFacilitator = currentUser.role === 'facilitator' && currentUser.site && currentUser.site !== 'All';

  for (let i = 1; i < rolloutData.length; i++) {
    const rolloutId = rolloutData[i][rolloutHeaders.indexOf('rolloutId')];
    const site = rolloutData[i][rolloutHeaders.indexOf('site')];
    if (isFacilitator && site !== currentUser.site) continue;

    const label = `${rolloutData[i][rolloutHeaders.indexOf('schoolName')]} (${rolloutData[i][rolloutHeaders.indexOf('period')]} ${rolloutData[i][rolloutHeaders.indexOf('year')]})`;
    rolloutMap[label.toLowerCase()] = {
      rolloutId: rolloutId,
      site: site,
      schoolName: rolloutData[i][rolloutHeaders.indexOf('schoolName')],
      period: rolloutData[i][rolloutHeaders.indexOf('period')],
      year: rolloutData[i][rolloutHeaders.indexOf('year')],
      status: rolloutData[i][rolloutHeaders.indexOf('status')],
      label: label
    };
  }

  if (Object.keys(rolloutMap).length === 0) {
    return { success: false, message: 'No accessible rollouts found for this user' };
  }

  // Participant map for upserts
  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const participantMap = {};
  for (let i = 1; i < pData.length; i++) {
    const pid = pData[i][pHeaders.indexOf('participantId')];
    participantMap[pid] = {
      rowIndex: i + 1,
      data: pData[i]
    };
  }

  const summary = {
    processed: 0,
    created: 0,
    updated: 0,
    skipped: 0,
    errors: []
  };

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rowNumber = r + 1; // 1-based CSV row

    const fullName = (row[idx('fullName')] || '').trim();
    const rolloutName = (row[idx('rolloutName')] || '').trim();

    if (!fullName && !rolloutName) {
      summary.skipped++;
      continue;
    }

    summary.processed++;

    if (!fullName) {
      summary.errors.push({ row: rowNumber, message: 'Full name is required' });
      summary.skipped++;
      continue;
    }
    if (!rolloutName) {
      summary.errors.push({ row: rowNumber, message: 'rolloutName is required' });
      summary.skipped++;
      continue;
    }
    const rollout = rolloutMap[rolloutName.toLowerCase()];
    if (!rollout) {
      summary.errors.push({ row: rowNumber, message: 'Rollout not found or not accessible: ' + rolloutName });
      summary.skipped++;
      continue;
    }

    // Create new participant (participantId auto-generated)
    if (isFacilitator && rollout.site !== currentUser.site) {
      summary.errors.push({ row: rowNumber, message: 'Unauthorized to enroll for site ' + rollout.site });
      summary.skipped++;
      continue;
    }

    const newParticipantId = generateParticipantId(rollout.site);
    const timestamp = new Date().toISOString();

    participantsSheet.appendRow([
      newParticipantId,
      fullName,
      rollout.site,
      rollout.rolloutId,
      rollout.schoolName,
      rollout.period,
      rollout.year,
      timestamp,
      currentUser.userId,
      'active',
      '',
      0
    ]);

    CONFIG.INSTRUMENTS.forEach(instrument => {
      const checklistId = generateUUID();
      checklistSheet.appendRow([
        checklistId,
        newParticipantId,
        instrument.number,
        instrument.name,
        instrument.category,
        'not_started',
        '',
        '',
        ''
      ]);
    });

    logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT_BULK', 'participant', newParticipantId,
      'Bulk enrolled: ' + fullName + ' into ' + rollout.schoolName);

    summary.created++;
  }

  return {
    success: true,
    summary: summary
  };
}

// ============================================
// CHECKLIST FUNCTIONS
// ============================================

/**
 * Update checklist item status
 */
function updateChecklistItem(token, checklistId, updateData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checklistSheet = ss.getSheetByName('Checklist');
  const participantsSheet = ss.getSheetByName('Participants');

  if (!checklistSheet || !participantsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const data = checklistSheet.getDataRange().getValues();
  const headers = data[0];

  // Build participant site map for authorization checks
  const participantSiteMap = {};
  if (participantsSheet) {
    const pData = participantsSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    for (let i = 1; i < pData.length; i++) {
      participantSiteMap[pData[i][pHeaders.indexOf('participantId')]] = pData[i][pHeaders.indexOf('site')];
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('checklistId')] === checklistId) {
      const participantId = data[i][headers.indexOf('participantId')];
      const instrumentName = data[i][headers.indexOf('instrumentName')];
      const participantSite = participantSiteMap[participantId];

      if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
        if (!participantSite) {
          return { success: false, message: 'Unable to verify participant site' };
        }
        if (participantSite !== currentUser.site) {
          return { success: false, message: 'Unauthorized for this site' };
        }
      }

      // Update status
      if (updateData.status) {
        checklistSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(updateData.status);

        // If marking as completed, set the date and user
        if (updateData.status === 'completed') {
          checklistSheet.getRange(i + 1, headers.indexOf('completedDate') + 1).setValue(new Date().toISOString());
          checklistSheet.getRange(i + 1, headers.indexOf('completedBy') + 1).setValue(currentUser.fullName);
        } else {
          // Clear completion info if status changed to not_started or missing
          checklistSheet.getRange(i + 1, headers.indexOf('completedDate') + 1).setValue('');
          checklistSheet.getRange(i + 1, headers.indexOf('completedBy') + 1).setValue('');
        }
      }

      // Update notes
      if (updateData.notes !== undefined) {
        checklistSheet.getRange(i + 1, headers.indexOf('notes') + 1).setValue(updateData.notes);
      }

      // Update participant's completion percentage
      updateParticipantCompletion(participantId);

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_CHECKLIST', 'checklist', checklistId,
        'Updated: ' + instrumentName + ' -> ' + updateData.status);

      return { success: true, message: 'Checklist item updated' };
    }
  }

  return { success: false, message: 'Checklist item not found' };
}

/**
 * Bulk update checklist items
 */
function bulkUpdateChecklist(token, participantId, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < updates.length; i++) {
    const update = updates[i];
    const result = updateChecklistItem(token, update.checklistId, {
      status: update.status,
      notes: update.notes
    });

    if (result.success) {
      successCount++;
    } else {
      errorCount++;
      if (result.message && result.message.toLowerCase().includes('unauthorized')) {
        return { success: false, message: result.message };
      }
    }
  }

  return {
    success: true,
    message: `Updated ${successCount} items, ${errorCount} errors`
  };
}

/**
 * Get checklist statuses for a specific instrument within a rollout
 */
function getInstrumentChecklistForRollout(token, rolloutId, instrumentNumber) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const rollout = getRolloutById(rolloutId);
  if (!rollout) {
    return { success: false, message: 'Rollout not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && rollout.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = CONFIG.INSTRUMENTS.find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const checklistData = checklistSheet.getDataRange().getValues();
  const cHeaders = checklistData[0];

  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const statusCol = cHeaders.indexOf('status');
  const checklistIdCol = cHeaders.indexOf('checklistId');

  const checklistMap = {};
  for (let i = 1; i < checklistData.length; i++) {
    if (String(checklistData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = checklistData[i][checklistParticipantCol];
      checklistMap[pid] = {
        checklistId: checklistData[i][checklistIdCol],
        status: checklistData[i][statusCol],
        completedDate: checklistData[i][cHeaders.indexOf('completedDate')],
        completedBy: checklistData[i][cHeaders.indexOf('completedBy')]
      };
    }
  }

  const participants = [];
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] !== rolloutId) continue;

    const participantId = pData[i][pHeaders.indexOf('participantId')];
    const checklistInfo = checklistMap[participantId] || {};

    participants.push({
      participantId: participantId,
      fullName: pData[i][pHeaders.indexOf('fullName')],
      site: pData[i][pHeaders.indexOf('site')],
      schoolName: pData[i][pHeaders.indexOf('schoolName')],
      status: checklistInfo.status || 'not_started',
      checklistId: checklistInfo.checklistId || '',
      completedDate: checklistInfo.completedDate || '',
      completedBy: checklistInfo.completedBy || ''
    });
  }

  return {
    success: true,
    instrumentName: instrument.name,
    rolloutName: rollout.schoolName,
    site: rollout.site,
    participants: participants
  };
}

/**
 * Bulk update an instrument across participants in a rollout
 */
function bulkUpdateInstrumentStatus(token, rolloutId, instrumentNumber, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const rollout = getRolloutById(rolloutId);
  if (!rollout) {
    return { success: false, message: 'Rollout not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && rollout.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = CONFIG.INSTRUMENTS.find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  if (!Array.isArray(updates) || updates.length === 0) {
    return { success: false, message: 'No updates provided' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const participantsInRollout = {};

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] === rolloutId) {
      participantsInRollout[pData[i][pHeaders.indexOf('participantId')]] = true;
    }
  }

  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];
  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const checklistIdCol = cHeaders.indexOf('checklistId');

  const checklistMap = {};
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = cData[i][checklistParticipantCol];
      checklistMap[pid] = cData[i][checklistIdCol];
    }
  }

  let successCount = 0;
  let errorCount = 0;

  updates.forEach(update => {
    if (!participantsInRollout[update.participantId]) {
      errorCount++;
      return;
    }

    if (CONFIG.CHECKLIST_STATUSES.indexOf(update.status) === -1) {
      errorCount++;
      return;
    }

    const checklistId = checklistMap[update.participantId];
    if (!checklistId) {
      errorCount++;
      return;
    }

    const result = updateChecklistItem(token, checklistId, { status: update.status });
    if (result.success) {
      successCount++;
    } else {
      errorCount++;
    }
  });

  logActivity(currentUser.userId, currentUser.fullName, 'BULK_UPDATE_INSTRUMENT', 'checklist', instrumentNumber,
    'Updated ' + successCount + ' ' + instrument.name + ' records for rollout ' + rollout.schoolName);

  return {
    success: errorCount === 0,
    message: 'Updated ' + successCount + ' participants' + (errorCount > 0 ? (' (' + errorCount + ' errors)') : '')
  };
}

/**
 * Update participant's completion percentage
 */
function updateParticipantCompletion(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  // Count completed items
  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];

  let total = 0;
  let completed = 0;

  for (let i = 1; i < cData.length; i++) {
    if (cData[i][cHeaders.indexOf('participantId')] === participantId) {
      total++;
      if (cData[i][cHeaders.indexOf('status')] === 'completed') {
        completed++;
      }
    }
  }

  const percentage = total > 0 ? Math.round((completed / total) * 100) : 0;

  // Update participant record
  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('participantId')] === participantId) {
      participantsSheet.getRange(i + 1, pHeaders.indexOf('completionPercentage') + 1).setValue(percentage);
      break;
    }
  }

  return percentage;
}

// ============================================
// DASHBOARD & STATISTICS FUNCTIONS
// ============================================

/**
 * Get dashboard statistics
 */
function getDashboardStats(token, siteFilter, rolloutFilter) {
  const currentUser = validateSession(token);
  // Allow public access for basic stats, but filter for authenticated users

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  const stats = {
    totalParticipants: 0,
    activeParticipants: 0,
    completedParticipants: 0,
    withdrawnParticipants: 0,
    ugaParticipants: 0,
    missouriParticipants: 0,
    activeRollouts: 0,
    overallCompletion: 0,
    instrumentStats: [],
    recentActivity: [],
    currentRollout: null
  };

  if (!participantsSheet) return stats;

  // Get participants data
  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];

  let totalCompletion = 0;
  let filteredParticipantIds = [];

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][pHeaders.indexOf('site')];
    const status = pData[i][pHeaders.indexOf('status')];
    const completion = pData[i][pHeaders.indexOf('completionPercentage')] || 0;
    const rolloutId = pData[i][pHeaders.indexOf('rolloutId')];
    const participantId = pData[i][pHeaders.indexOf('participantId')];

    // Apply site filter
    if (siteFilter && siteFilter !== 'All' && site !== siteFilter) {
      continue;
    }

    // Apply rollout filter
    if (rolloutFilter && rolloutFilter !== 'All' && rolloutId !== rolloutFilter) {
      continue;
    }

    filteredParticipantIds.push(participantId);
    stats.totalParticipants++;
    totalCompletion += completion;

    if (status === 'active') stats.activeParticipants++;
    if (status === 'completed') stats.completedParticipants++;
    if (status === 'withdrawn') stats.withdrawnParticipants++;

    if (site === 'UGA') stats.ugaParticipants++;
    if (site === 'Missouri') stats.missouriParticipants++;
  }

  stats.overallCompletion = stats.totalParticipants > 0
    ? Math.round(totalCompletion / stats.totalParticipants)
    : 0;

  // Get rollouts data
  if (rolloutsSheet) {
    const rData = rolloutsSheet.getDataRange().getValues();
    const rHeaders = rData[0];

    for (let i = 1; i < rData.length; i++) {
      const rSite = rData[i][rHeaders.indexOf('site')];
      const rStatus = rData[i][rHeaders.indexOf('status')];
      const rId = rData[i][rHeaders.indexOf('rolloutId')];

      // If a specific rollout is selected, get its details
      if (rolloutFilter && rolloutFilter !== 'All' && rId === rolloutFilter) {
        stats.currentRollout = {
          rolloutId: rId,
          site: rSite,
          schoolName: rData[i][rHeaders.indexOf('schoolName')],
          period: rData[i][rHeaders.indexOf('period')],
          year: rData[i][rHeaders.indexOf('year')],
          status: rStatus
        };
      }

      if (siteFilter && siteFilter !== 'All' && rSite !== siteFilter) {
        continue;
      }

      if (rStatus === 'active') stats.activeRollouts++;
    }
  }

  // Get instrument completion stats (only for filtered participants)
  if (checklistSheet) {
    const cData = checklistSheet.getDataRange().getValues();
    const cHeaders = cData[0];

    const instrumentCounts = {};

    CONFIG.INSTRUMENTS.forEach(inst => {
      instrumentCounts[inst.number] = { total: 0, completed: 0, name: inst.name };
    });

    for (let i = 1; i < cData.length; i++) {
      const instNum = cData[i][cHeaders.indexOf('instrumentNumber')];
      const instStatus = cData[i][cHeaders.indexOf('status')];
      const checklistParticipantId = cData[i][cHeaders.indexOf('participantId')];

      // Only count checklist items for filtered participants
      if (filteredParticipantIds.length > 0 && !filteredParticipantIds.includes(checklistParticipantId)) {
        continue;
      }

      if (instrumentCounts[instNum]) {
        instrumentCounts[instNum].total++;
        if (instStatus === 'completed') {
          instrumentCounts[instNum].completed++;
        }
      }
    }

    stats.instrumentStats = CONFIG.INSTRUMENTS.map(inst => ({
      number: inst.number,
      name: inst.name,
      category: inst.category,
      total: instrumentCounts[inst.number].total,
      completed: instrumentCounts[inst.number].completed,
      percentage: instrumentCounts[inst.number].total > 0
        ? Math.round((instrumentCounts[inst.number].completed / instrumentCounts[inst.number].total) * 100)
        : 0
    }));
  }

  return stats;
}

/**
 * Get analytics overview across rollouts and instruments
 */
function getAnalyticsOverview(token, siteFilter, rolloutFilter) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const effectiveSite = (currentUser.role === 'facilitator' && currentUser.site !== 'All')
    ? currentUser.site
    : (siteFilter || 'All');
  const effectiveRollout = rolloutFilter || 'All';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet) {
    return {
      success: true,
      overview: {
        totalParticipants: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        averageCompletion: 0,
        activeRollouts: 0,
        rolloutCount: 0,
        instrumentsCompleted: 0,
        instrumentsTotal: 0,
        siteFilter: effectiveSite,
        rolloutFilter: effectiveRollout
      },
      rollouts: [],
      instrumentStats: []
    };
  }

  const rolloutEntries = {};
  if (rolloutsSheet) {
    const rData = rolloutsSheet.getDataRange().getValues();
    const rHeaders = rData[0];

    for (let i = 1; i < rData.length; i++) {
      const rolloutId = rData[i][rHeaders.indexOf('rolloutId')];
      const rolloutSite = rData[i][rHeaders.indexOf('site')];

      if (effectiveSite !== 'All' && rolloutSite !== effectiveSite) continue;
      if (effectiveRollout !== 'All' && rolloutId !== effectiveRollout) continue;

      rolloutEntries[rolloutId] = {
        rolloutId: rolloutId,
        site: rolloutSite,
        schoolName: rData[i][rHeaders.indexOf('schoolName')],
        period: rData[i][rHeaders.indexOf('period')],
        year: rData[i][rHeaders.indexOf('year')],
        status: rData[i][rHeaders.indexOf('status')],
        participantCount: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        completionSum: 0,
        averageCompletion: 0
      };
    }
  }

  const overview = {
    totalParticipants: 0,
    activeParticipants: 0,
    completedParticipants: 0,
    withdrawnParticipants: 0,
    averageCompletion: 0,
    activeRollouts: 0,
    rolloutCount: Object.keys(rolloutEntries).length,
    instrumentsCompleted: 0,
    instrumentsTotal: 0,
    siteFilter: effectiveSite,
    rolloutFilter: effectiveRollout
  };

  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const includedParticipantIds = [];
  let completionSum = 0;

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][pHeaders.indexOf('site')];
    const rolloutId = pData[i][pHeaders.indexOf('rolloutId')];
    const status = pData[i][pHeaders.indexOf('status')];
    const completion = Number(pData[i][pHeaders.indexOf('completionPercentage')]) || 0;
    const participantId = pData[i][pHeaders.indexOf('participantId')];

    if (effectiveSite !== 'All' && site !== effectiveSite) continue;
    if (effectiveRollout !== 'All' && rolloutId !== effectiveRollout) continue;

    includedParticipantIds.push(participantId);
    overview.totalParticipants++;
    overview.activeParticipants += status === 'active' ? 1 : 0;
    overview.completedParticipants += status === 'completed' ? 1 : 0;
    overview.withdrawnParticipants += status === 'withdrawn' ? 1 : 0;
    completionSum += completion;

    const key = rolloutId || 'Unassigned';
    if (!rolloutEntries[key]) {
      rolloutEntries[key] = {
        rolloutId: rolloutId || 'Unassigned',
        site: site,
        schoolName: rolloutId ? pData[i][pHeaders.indexOf('schoolName')] : 'No rollout assigned',
        period: pData[i][pHeaders.indexOf('period')],
        year: pData[i][pHeaders.indexOf('year')],
        status: rolloutId ? 'active' : 'unassigned',
        participantCount: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        completionSum: 0,
        averageCompletion: 0
      };
    }

    const rolloutStats = rolloutEntries[key];
    rolloutStats.participantCount++;
    rolloutStats.activeParticipants += status === 'active' ? 1 : 0;
    rolloutStats.completedParticipants += status === 'completed' ? 1 : 0;
    rolloutStats.withdrawnParticipants += status === 'withdrawn' ? 1 : 0;
    rolloutStats.completionSum += completion;
  }

  overview.averageCompletion = overview.totalParticipants > 0
    ? Math.round(completionSum / overview.totalParticipants)
    : 0;

  Object.keys(rolloutEntries).forEach(key => {
    const entry = rolloutEntries[key];
    if (entry.participantCount > 0) {
      entry.averageCompletion = Math.round(entry.completionSum / entry.participantCount);
    }
    if (entry.status === 'active') {
      overview.activeRollouts++;
    }
  });

  // Instrument progress for included participants
  let instrumentStats = [];
  if (checklistSheet && includedParticipantIds.length > 0) {
    const cData = checklistSheet.getDataRange().getValues();
    const cHeaders = cData[0];
    const counts = {};

    CONFIG.INSTRUMENTS.forEach(inst => {
      counts[inst.number] = { total: 0, completed: 0, name: inst.name, category: inst.category };
    });

    for (let i = 1; i < cData.length; i++) {
      const participantId = cData[i][cHeaders.indexOf('participantId')];
      if (!includedParticipantIds.includes(participantId)) continue;

      const instrumentNumber = cData[i][cHeaders.indexOf('instrumentNumber')];
      const status = cData[i][cHeaders.indexOf('status')];

      if (counts[instrumentNumber]) {
        counts[instrumentNumber].total++;
        if (status === 'completed') {
          counts[instrumentNumber].completed++;
        }
      }
    }

    instrumentStats = CONFIG.INSTRUMENTS.map(inst => ({
      number: inst.number,
      name: inst.name,
      category: inst.category,
      total: counts[inst.number].total,
      completed: counts[inst.number].completed,
      percentage: counts[inst.number].total > 0
        ? Math.round((counts[inst.number].completed / counts[inst.number].total) * 100)
        : 0
    }));

    overview.instrumentsCompleted = instrumentStats.reduce((sum, inst) => sum + inst.completed, 0);
    overview.instrumentsTotal = instrumentStats.reduce((sum, inst) => sum + inst.total, 0);
  }

  const rollouts = Object.values(rolloutEntries).sort((a, b) => {
    if (a.site === b.site) {
      return (b.year || '').toString().localeCompare((a.year || '').toString());
    }
    return a.site.localeCompare(b.site);
  });

  overview.rolloutCount = rollouts.length;

  return {
    success: true,
    overview: overview,
    rollouts: rollouts,
    instrumentStats: instrumentStats
  };
}

/**
 * Get public dashboard statistics (no authentication required)
 */
function getPublicDashboardStats(siteFilter, rolloutFilter) {
  return getDashboardStats(null, siteFilter || 'All', rolloutFilter || 'All');
}

/**
 * Get public rollouts list (no authentication required)
 */
function getPublicRollouts(siteFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return { success: true, rollouts: [] };

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const rollouts = [];

  for (let i = 1; i < data.length; i++) {
    const site = data[i][headers.indexOf('site')];
    const status = data[i][headers.indexOf('status')];

    // Apply site filter
    if (siteFilter && siteFilter !== 'All' && site !== siteFilter) {
      continue;
    }

    // Only return active rollouts for public view
    if (status === 'active') {
      rollouts.push({
        rolloutId: data[i][headers.indexOf('rolloutId')],
        site: site,
        schoolName: data[i][headers.indexOf('schoolName')],
        period: data[i][headers.indexOf('period')],
        year: data[i][headers.indexOf('year')],
        status: status
      });
    }
  }

  return { success: true, rollouts: rollouts };
}

/**
 * Get site comparison data
 */
function getSiteComparison(token) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ugaStats = getDashboardStats(token, 'UGA', 'All');
  const missouriStats = getDashboardStats(token, 'Missouri', 'All');

  return {
    success: true,
    comparison: {
      UGA: ugaStats,
      Missouri: missouriStats
    }
  };
}

// ============================================
// ACTIVITY LOG FUNCTIONS
// ============================================

/**
 * Log an activity
 */
function logActivity(userId, userName, action, targetType, targetId, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ActivityLog');

  if (!logSheet) return;

  const logId = generateUUID();
  const timestamp = new Date().toISOString();

  logSheet.appendRow([
    logId,
    timestamp,
    userId,
    userName,
    action,
    targetType,
    targetId,
    details
  ]);
}

/**
 * Get recent activity logs
 */
function getRecentActivity(token, limit) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  limit = limit || 50;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ActivityLog');

  if (!logSheet) return { success: true, activities: [] };

  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  const activities = [];

  // Get last 'limit' entries (reverse order)
  for (let i = Math.min(data.length - 1, limit); i >= 1; i--) {
    activities.push({
      logId: data[i][headers.indexOf('logId')],
      timestamp: data[i][headers.indexOf('timestamp')],
      userId: data[i][headers.indexOf('userId')],
      userName: data[i][headers.indexOf('userName')],
      action: data[i][headers.indexOf('action')],
      targetType: data[i][headers.indexOf('targetType')],
      targetId: data[i][headers.indexOf('targetId')],
      details: data[i][headers.indexOf('details')]
    });
  }

  // Sort by timestamp descending
  activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  return { success: true, activities: activities.slice(0, limit) };
}

// ============================================
// EXPORT FUNCTIONS
// ============================================

/**
 * Export participants data to CSV format
 */
function exportParticipantsCSV(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  // Enforce facilitator site scope
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    filters = filters || {};
    filters.site = currentUser.site;
  }

  const result = getAllParticipants(token, filters);
  if (!result.success) return result;

  const participants = result.participants;

  // Create CSV content
  const headers = ['Participant ID', 'Full Name', 'Site', 'School', 'Period', 'Year',
    'Enrollment Date', 'Status', 'Completion %', 'Notes'];

  let csv = headers.join(',') + '\n';

  participants.forEach(p => {
    const row = [
      p.participantId,
      '"' + (p.fullName || '').replace(/"/g, '""') + '"',
      p.site,
      '"' + (p.schoolName || '').replace(/"/g, '""') + '"',
      p.period,
      p.year,
      p.enrollmentDate,
      p.status,
      p.completionPercentage,
      '"' + (p.notes || '').replace(/"/g, '""') + '"'
    ];
    csv += row.join(',') + '\n';
  });

  return { success: true, csv: csv, filename: 'participants_export_' + new Date().toISOString().split('T')[0] + '.csv' };
}

/**
 * Export full checklist data
 */
function exportChecklistCSV(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  // Enforce facilitator site scope
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    filters = filters || {};
    filters.site = currentUser.site;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Get participant data for reference
  const pData = participantsSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const participantMap = {};

  for (let i = 1; i < pData.length; i++) {
    const pid = pData[i][pHeaders.indexOf('participantId')];
    participantMap[pid] = {
      fullName: pData[i][pHeaders.indexOf('fullName')],
      site: pData[i][pHeaders.indexOf('site')],
      schoolName: pData[i][pHeaders.indexOf('schoolName')],
      rolloutId: pData[i][pHeaders.indexOf('rolloutId')]
    };
  }

  // Get checklist data
  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];

  const headers = ['Participant ID', 'Participant Name', 'Site', 'School',
    'Instrument #', 'Instrument Name', 'Category', 'Status', 'Completed Date', 'Completed By', 'Notes'];

  let csv = headers.join(',') + '\n';

  for (let i = 1; i < cData.length; i++) {
    const pid = cData[i][cHeaders.indexOf('participantId')];
    const participant = participantMap[pid] || {};

    // Apply filters
    if (filters && filters.site && filters.site !== 'All' && participant.site !== filters.site) {
      continue;
    }
    if (filters && filters.rolloutId && filters.rolloutId !== 'All' && participant.rolloutId !== filters.rolloutId) {
      continue;
    }

    const row = [
      pid,
      '"' + (participant.fullName || '').replace(/"/g, '""') + '"',
      participant.site || '',
      '"' + (participant.schoolName || '').replace(/"/g, '""') + '"',
      cData[i][cHeaders.indexOf('instrumentNumber')],
      '"' + (cData[i][cHeaders.indexOf('instrumentName')] || '').replace(/"/g, '""') + '"',
      cData[i][cHeaders.indexOf('category')],
      cData[i][cHeaders.indexOf('status')],
      cData[i][cHeaders.indexOf('completedDate')],
      '"' + (cData[i][cHeaders.indexOf('completedBy')] || '').replace(/"/g, '""') + '"',
      '"' + (cData[i][cHeaders.indexOf('notes')] || '').replace(/"/g, '""') + '"'
    ];
    csv += row.join(',') + '\n';
  }

  return { success: true, csv: csv, filename: 'checklist_export_' + new Date().toISOString().split('T')[0] + '.csv' };
}

/**
 * Export summary report
 */
function exportSummaryReport(token) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const stats = getDashboardStats(token, 'All');
  const ugaStats = getDashboardStats(token, 'UGA');
  const missouriStats = getDashboardStats(token, 'Missouri');

  let report = 'G4G Research Operations Summary Report\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n\n';

  report += '=== OVERALL STATISTICS ===\n';
  report += 'Total Participants: ' + stats.totalParticipants + '\n';
  report += 'Active Participants: ' + stats.activeParticipants + '\n';
  report += 'Completed: ' + stats.completedParticipants + '\n';
  report += 'Withdrawn: ' + stats.withdrawnParticipants + '\n';
  report += 'Overall Completion: ' + stats.overallCompletion + '%\n\n';

  report += '=== UGA STATISTICS ===\n';
  report += 'Total Participants: ' + ugaStats.totalParticipants + '\n';
  report += 'Active: ' + ugaStats.activeParticipants + '\n';
  report += 'Completion: ' + ugaStats.overallCompletion + '%\n\n';

  report += '=== MISSOURI STATISTICS ===\n';
  report += 'Total Participants: ' + missouriStats.totalParticipants + '\n';
  report += 'Active: ' + missouriStats.activeParticipants + '\n';
  report += 'Completion: ' + missouriStats.overallCompletion + '%\n\n';

  report += '=== INSTRUMENT COMPLETION ===\n';
  stats.instrumentStats.forEach(inst => {
    report += inst.number + '. ' + inst.name + ': ' + inst.percentage + '% (' + inst.completed + '/' + inst.total + ')\n';
  });

  return { success: true, report: report, filename: 'summary_report_' + new Date().toISOString().split('T')[0] + '.txt' };
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Generate a UUID
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Generate a random password
 */
function generateRandomPassword() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#$%';
  let password = '';
  for (let i = 0; i < 12; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

/**
 * Generate a session token
 */
function generateSessionToken() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

/**
 * Escape a value for CSV
 */
function csvEscape(value) {
  const str = value === null || value === undefined ? '' : String(value);
  return '"' + str.replace(/"/g, '""') + '"';
}

/**
 * Hash a password using SHA-256
 */
function hashPassword(password) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return hash.map(byte => {
    const hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/**
 * Verify a password against stored password
 * Using plain text comparison since the Google Sheet is highly protected
 */
function verifyPassword(password, storedPassword) {
  // Plain text comparison - storedPassword is stored as-is in the sheet
  return password === storedPassword;
}

/**
 * Format date for display
 */
function formatDate(dateString) {
  if (!dateString) return '';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
}

/**
 * Format date and time for display
 */
function formatDateTime(dateString) {
  if (!dateString) return '';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}
