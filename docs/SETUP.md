# G4G Web App - Setup Guide

## Prerequisites

- Google Account with access to Google Drive
- Google Sheets
- Google Apps Script

## Step-by-Step Deployment

### Step 1: Create Google Spreadsheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it "G4G Research Operations Database"
4. Note the Spreadsheet ID from the URL:
   - URL format: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`
   - Copy the `SPREADSHEET_ID` portion

### Step 2: Open Apps Script Editor

1. In your spreadsheet, go to **Extensions > Apps Script**
2. This opens the Apps Script editor
3. Delete any existing code in `Code.gs`

### Step 3: Create Script Files

Create the following `.gs` files in the Apps Script editor:

1. **Code.gs** - Copy content from `src/Code.gs`
2. **Database.gs** - Copy content from `src/Database.gs`
3. **Auth.gs** - Copy content from `src/Auth.gs`
4. **Utils.gs** - Copy content from `src/Utils.gs`

To create a new file:
- Click the `+` button next to "Files"
- Select "Script"
- Name it appropriately (without .gs extension)

### Step 4: Create HTML Files

Create the following HTML files:

1. **index.html** - Copy from `src/html/index.html`
2. **login.html** - Copy from `src/html/login.html`
3. **dashboard.html** - Copy from `src/html/dashboard.html`
4. **users.html** - Copy from `src/html/users.html`
5. **rollouts.html** - Copy from `src/html/rollouts.html`
6. **participants.html** - Copy from `src/html/participants.html`
7. **participant-detail.html** - Copy from `src/html/participant-detail.html`
8. **reports.html** - Copy from `src/html/reports.html`

To create a new HTML file:
- Click the `+` button next to "Files"
- Select "HTML"
- Name it appropriately (without .html extension)

### Step 5: Initialize Database

1. In the Apps Script editor, select `Code.gs`
2. In the function dropdown (top toolbar), select `initializeDatabase`
3. Click **Run**
4. Grant permissions when prompted:
   - Review permissions
   - Click "Advanced"
   - Click "Go to G4G Research Operations (unsafe)"
   - Click "Allow"
5. Check your spreadsheet - you should see new sheets created:
   - Users
   - StudyRollouts
   - Participants
   - Checklist
   - ActivityLog
   - Config

### Step 6: Create First Admin User

1. In the Apps Script editor, select `Code.gs`
2. In the function dropdown, select `createInitialAdmin`
3. Click **Run**
4. Check the execution log for the temporary password
5. **Important**: Save this password - you'll need it for first login

Default admin credentials:
- Email: `admin@g4g.edu`
- Password: (shown in execution log)

### Step 7: Deploy as Web App

1. Click **Deploy > New deployment**
2. Click the gear icon next to "Select type"
3. Select **Web app**
4. Configure:
   - Description: "G4G Research Operations v1.0"
   - Execute as: **Me**
   - Who has access: **Anyone** (or "Anyone within [your organization]" for restricted access)
5. Click **Deploy**
6. Copy the Web App URL
7. Click **Done**

### Step 8: Test the Application

1. Open the Web App URL in a browser
2. You should see the public dashboard
3. Click "Login"
4. Enter admin credentials
5. Verify you have full admin access

## Post-Deployment Configuration

### Add Additional Admin Users

1. Log in as admin
2. Go to Users section
3. Click "Add User"
4. Fill in details and select "Admin" role

### Set Up Study Cohorts

Before enrolling participants:
1. Go to Study Cohorts section
2. Click "Add Cohort"
3. Configure:
   - Site (UGA or Missouri)
   - School/Location name
   - Period (Spring/Summer/Fall)
   - Year

### Create Facilitator Accounts

For each research staff member who will be entering data:
1. Go to Users section
2. Click "Add User"
3. Select "Facilitator" role
4. Assign to appropriate site

## Updating the Application

To update after code changes:

1. Open Apps Script editor
2. Make necessary changes
3. Click **Deploy > Manage deployments**
4. Click the pencil icon to edit
5. Update version to "New version"
6. Click **Deploy**

## Troubleshooting

### "Authorization Required" Error
- Re-run `initializeDatabase()` and grant permissions

### Sheets Not Created
- Ensure you have edit access to the spreadsheet
- Check execution log for errors

### Login Not Working
- Verify the Users sheet has data
- Check that password was properly hashed
- Try running `resetAdminPassword()` function

### Web App Shows Error
- Check execution logs: View > Executions
- Verify all HTML files are created correctly
- Ensure all .gs files are saved

## Security Notes

1. **Change default admin password** immediately after first login
2. Use strong passwords for all accounts
3. Consider restricting web app access to your organization
4. Regularly review activity logs
5. Remove inactive user accounts promptly

## Backup Recommendations

1. Regularly export data using the Reports feature
2. Make periodic copies of the Google Spreadsheet
3. Keep version history enabled in Google Sheets
