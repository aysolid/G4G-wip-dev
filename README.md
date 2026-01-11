# Gaming4Good (G4G) Research Operations Web App

A cross-site research operations web application built with Google Apps Script and Google Sheets for tracking participant progress in the Gaming4Good multi-site research study.

## Overview

G4G is a grant-funded project that engages neurodiverse learners in video game design activities using Game Builder Garage software on Nintendo Switch. This application serves as a centralized tracking system for participant data collection across two research sites: University of Georgia (UGA) and University of Missouri (Mizzou).

## Features

- **Participant Tracking**: Track each participant from enrollment through post-study completion
- **18-Item Protocol Checklist**: Monitor completion of all required research instruments
- **Role-Based Access**: Admin, Facilitator, and Viewer roles with appropriate permissions
- **Cross-Site Monitoring**: View and compare progress across UGA and Missouri sites
- **Study Cohort Management**: Organize participants by cohorts/sessions
- **Real-Time Dashboards**: Visualize completion rates and identify missing data
- **Export Functionality**: Generate reports for analysis and compliance

## The 18 Non-Negotiable Protocol Items

1. Consent Form
2. Assent Form / Pre-test
3. Lesson 1 Journal
4. Lesson 2 Journal
5. Lesson 3 Journal
6. Lesson 3 - Design a Game Worksheet
7. Lesson 4 Journal
8. Lesson 4 - Paper Prototyping Worksheet
9. Lesson 5 Journal
10. Lesson 5 - Debugging Worksheet
11. Lesson 6 Journal
12. Lesson 6 - Game Refinement Worksheet
13. Lesson 7 Journal
14. Lesson 7 - Playtesting Feedback Guide
15. Lesson 8 Journal
16. Post-Test
17. Participant Feedback Survey
18. Parent Satisfaction Survey

## User Roles & Permissions

### Admin
- Full access to all features
- Create/manage users and assign roles
- Set up study rollouts
- Enroll participants
- Update checklist items
- View all sites
- Export data

### Facilitator
- Enroll participants
- Update checklist items
- Add notes
- View participants (site-filtered by default)

### Viewer
- View-only access to dashboards
- View participant details
- Cannot modify any data

## Technical Stack

- **Database**: Google Sheets
- **Backend**: Google Apps Script
- **Frontend**: HTML, CSS, JavaScript (embedded per Apps Script requirements)
- **Deployment**: Google Apps Script Web App

## Project Structure

```
G4G-wip-dev/
├── README.md
├── docs/
│   ├── SETUP.md              # Deployment instructions
│   ├── USER_GUIDE.md         # User documentation
│   └── DATABASE_SCHEMA.md    # Data structure documentation
└── src/
    ├── Code.gs               # Main backend logic
    ├── Database.gs           # Database operations
    ├── Auth.gs               # Authentication functions
    ├── Utils.gs              # Utility functions
    └── html/
        ├── index.html        # Public dashboard
        ├── login.html        # Login page
        ├── dashboard.html    # Authenticated dashboard
        ├── users.html        # User management (Admin)
        ├── rollouts.html     # Study rollout management (Admin)
        ├── participants.html # Participant list
        ├── participant-detail.html # Individual participant view
        ├── reports.html      # Reports and exports
        └── includes/
            ├── styles.html   # Shared CSS
            └── scripts.html  # Shared JavaScript
```

## Google Sheets Database Structure

### Sheet: Users
| Column | Description |
|--------|-------------|
| userId | Unique identifier |
| email | User email (login) |
| passwordHash | Hashed password |
| fullName | Display name |
| role | admin/facilitator/viewer |
| site | UGA/Missouri/All |
| status | active/inactive |
| createdAt | Timestamp |
| createdBy | Admin userId |

### Sheet: StudyRollouts
| Column | Description |
|--------|-------------|
| rolloutId | Unique identifier |
| site | UGA/Missouri |
| schoolName | Name of school/location |
| period | Spring/Summer/Fall |
| year | Year (e.g., 2025) |
| status | active/completed/cancelled |
| createdAt | Timestamp |
| createdBy | Admin userId |

### Sheet: Participants
| Column | Description |
|--------|-------------|
| participantId | Auto-generated ID |
| fullName | Participant's full name |
| site | UGA/Missouri |
| rolloutId | Associated rollout |
| schoolName | From rollout |
| period | From rollout |
| year | From rollout |
| enrollmentDate | Auto-generated |
| enrolledBy | User who enrolled |
| status | active/withdrawn/completed |
| notes | Additional notes |

### Sheet: Checklist
| Column | Description |
|--------|-------------|
| checklistId | Unique identifier |
| participantId | Associated participant |
| instrumentNumber | 1-18 |
| instrumentName | Name of instrument |
| status | not_started/completed/missing |
| completedDate | When marked complete |
| completedBy | User who marked it |
| notes | Additional notes |

### Sheet: ActivityLog
| Column | Description |
|--------|-------------|
| logId | Unique identifier |
| timestamp | When action occurred |
| userId | Who performed action |
| action | Type of action |
| targetType | participant/checklist/user/rollout |
| targetId | ID of affected record |
| details | JSON with details |

## Setup Instructions

See [SETUP.md](docs/SETUP.md) for detailed deployment instructions.

## Quick Start

1. Create a new Google Spreadsheet
2. Open Script Editor (Extensions > Apps Script)
3. Copy all `.gs` files from `src/` to the script editor
4. Create HTML files in the script editor matching `src/html/`
5. Run the `initializeDatabase()` function to set up sheets
6. Deploy as Web App
7. Create the first admin user through the setup wizard

## License

This project is developed for research purposes by the Gaming4Good research team.

## Contact

For questions or support, contact the G4G research team.
