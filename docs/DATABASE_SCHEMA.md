# G4G Research Operations - Database Schema

This document describes the Google Sheets database structure used by the G4G Research Operations application.

## Overview

The application uses a single Google Spreadsheet with multiple sheets (tabs) serving as database tables. Each sheet has a header row with column names, followed by data rows.

---

## Sheet: Users

Stores user accounts and authentication information.

| Column | Type | Description |
|--------|------|-------------|
| userId | String (UUID) | Unique identifier for the user |
| email | String | User's email address (used for login) |
| passwordHash | String | SHA-256 hash of the user's password |
| fullName | String | User's display name |
| role | Enum | User role: `admin`, `facilitator`, `viewer` |
| site | String | Site assignment: `All`, `UGA`, `Missouri` |
| status | Enum | Account status: `active`, `inactive` |
| createdAt | ISO DateTime | When the account was created |
| createdBy | String | userId of the admin who created this account |
| lastLogin | ISO DateTime | Timestamp of last successful login |

### Example Row
```
userId: a1b2c3d4-e5f6-...
email: researcher@uga.edu
passwordHash: 5e884898da28047d...
fullName: Jane Smith
role: facilitator
site: UGA
status: active
createdAt: 2025-01-15T10:30:00Z
createdBy: admin-uuid-here
lastLogin: 2025-01-20T14:22:00Z
```

---

## Sheet: StudyRollouts

Stores study cohort/rollout configurations.

| Column | Type | Description |
|--------|------|-------------|
| rolloutId | String (UUID) | Unique identifier for the rollout |
| site | Enum | Research site: `UGA`, `Missouri` |
| schoolName | String | Name of the school or location |
| period | Enum | Academic period: `Spring`, `Summer`, `Fall` |
| year | Number | Year (e.g., 2025) |
| status | Enum | Rollout status: `active`, `completed`, `cancelled` |
| createdAt | ISO DateTime | When the rollout was created |
| createdBy | String | userId of the admin who created this rollout |
| description | String | Optional description of the cohort |

### Example Row
```
rolloutId: r1o2l3l4-o5u6-...
site: UGA
schoolName: Double Helix School
period: Fall
year: 2025
status: active
createdAt: 2025-01-10T09:00:00Z
createdBy: admin-uuid-here
description: First cohort for Fall 2025 semester
```

---

## Sheet: Participants

Stores enrolled participant information.

| Column | Type | Description |
|--------|------|-------------|
| participantId | String | Auto-generated ID with site prefix (e.g., UGA-ABC123) |
| fullName | String | Participant's full name |
| site | Enum | Research site: `UGA`, `Missouri` |
| rolloutId | String | Reference to StudyRollouts.rolloutId |
| schoolName | String | Copied from rollout for quick reference |
| period | String | Copied from rollout for quick reference |
| year | Number | Copied from rollout for quick reference |
| enrollmentDate | ISO DateTime | When the participant was enrolled |
| enrolledBy | String | userId of the user who enrolled this participant |
| status | Enum | Participant status: `active`, `withdrawn`, `completed` |
| notes | String | Optional notes about the participant |
| completionPercentage | Number | Calculated % of checklist items completed (0-100) |

### Participant ID Format
- UGA participants: `UGA-{timestamp}{random}` (e.g., UGA-LXYZAB12)
- Missouri participants: `MIZ-{timestamp}{random}` (e.g., MIZ-LXYZAB12)

### Example Row
```
participantId: UGA-LXYZ1234
fullName: John Doe
site: UGA
rolloutId: r1o2l3l4-o5u6-...
schoolName: Double Helix School
period: Fall
year: 2025
enrollmentDate: 2025-01-15T10:30:00Z
enrolledBy: facilitator-uuid-here
status: active
notes: Requires additional support
completionPercentage: 44
```

---

## Sheet: Checklist

Stores the 18-item protocol checklist for each participant.

| Column | Type | Description |
|--------|------|-------------|
| checklistId | String (UUID) | Unique identifier for this checklist item |
| participantId | String | Reference to Participants.participantId |
| instrumentNumber | Number | Instrument sequence number (1-18) |
| instrumentName | String | Name of the instrument |
| category | Enum | Category: `enrollment`, `lesson`, `worksheet`, `post` |
| status | Enum | Status: `not_started`, `completed`, `missing` |
| completedDate | ISO DateTime | When the item was marked complete |
| completedBy | String | Name of the user who marked it complete |
| notes | String | Optional notes about this specific item |

### The 18 Instruments

| # | Name | Category |
|---|------|----------|
| 1 | Consent Form | enrollment |
| 2 | Assent Form / Pre-test | enrollment |
| 3 | Lesson 1 Journal | lesson |
| 4 | Lesson 2 Journal | lesson |
| 5 | Lesson 3 Journal | lesson |
| 6 | Lesson 3 - Design a Game Worksheet | worksheet |
| 7 | Lesson 4 Journal | lesson |
| 8 | Lesson 4 - Paper Prototyping Worksheet | worksheet |
| 9 | Lesson 5 Journal | lesson |
| 10 | Lesson 5 - Debugging Worksheet | worksheet |
| 11 | Lesson 6 Journal | lesson |
| 12 | Lesson 6 - Game Refinement Worksheet | worksheet |
| 13 | Lesson 7 Journal | lesson |
| 14 | Lesson 7 - Playtesting Feedback Guide | worksheet |
| 15 | Lesson 8 Journal | lesson |
| 16 | Post-Test | post |
| 17 | Participant Feedback Survey | post |
| 18 | Parent Satisfaction Survey | post |

### Example Row
```
checklistId: c1h2e3c4-k5l6-...
participantId: UGA-LXYZ1234
instrumentNumber: 3
instrumentName: Lesson 1 Journal
category: lesson
status: completed
completedDate: 2025-01-16T11:45:00Z
completedBy: Jane Smith
notes: Completed during session 1
```

---

## Sheet: Sessions

Stores active user sessions for authentication.

| Column | Type | Description |
|--------|------|-------------|
| sessionId | String (UUID) | Unique identifier for the session |
| userId | String | Reference to Users.userId |
| token | String | Session token (UUID-UUID format) |
| createdAt | ISO DateTime | When the session was created |
| expiresAt | ISO DateTime | When the session expires (24 hours after creation) |
| isActive | Boolean | Whether the session is currently active |

### Example Row
```
sessionId: s1e2s3s4-i5o6-...
userId: a1b2c3d4-e5f6-...
token: abc123-def456-...-ghi789-jkl012-...
createdAt: 2025-01-20T14:22:00Z
expiresAt: 2025-01-21T14:22:00Z
isActive: TRUE
```

---

## Sheet: ActivityLog

Stores audit trail of all actions in the system.

| Column | Type | Description |
|--------|------|-------------|
| logId | String (UUID) | Unique identifier for the log entry |
| timestamp | ISO DateTime | When the action occurred |
| userId | String | Reference to Users.userId |
| userName | String | User's display name for quick reference |
| action | String | Type of action performed |
| targetType | Enum | Type of entity affected: `user`, `rollout`, `participant`, `checklist` |
| targetId | String | ID of the affected entity |
| details | String | Human-readable description of the action |

### Action Types
- `LOGIN` - User logged in
- `LOGOUT` - User logged out
- `CREATE_USER` - New user created
- `UPDATE_USER` - User information updated
- `RESET_PASSWORD` - User password reset
- `PASSWORD_CHANGE` - User changed their own password
- `CREATE_ROLLOUT` - New rollout created
- `UPDATE_ROLLOUT` - Rollout updated
- `ENROLL_PARTICIPANT` - New participant enrolled
- `UPDATE_PARTICIPANT` - Participant info updated
- `UPDATE_CHECKLIST` - Checklist item status changed

### Example Row
```
logId: l1o2g3i4-d5e6-...
timestamp: 2025-01-20T14:30:00Z
userId: a1b2c3d4-e5f6-...
userName: Jane Smith
action: UPDATE_CHECKLIST
targetType: checklist
targetId: c1h2e3c4-k5l6-...
details: Updated: Lesson 1 Journal -> completed
```

---

## Sheet: Config

Stores application configuration (reserved for future use).

| Column | Type | Description |
|--------|------|-------------|
| key | String | Configuration key |
| value | String | Configuration value |
| description | String | Description of the setting |
| updatedAt | ISO DateTime | When the setting was last updated |

---

## Data Relationships

```
Users
  ├── Creates → StudyRollouts (createdBy)
  ├── Creates → Participants (enrolledBy)
  ├── Creates → Users (createdBy)
  └── Has → Sessions (userId)

StudyRollouts
  └── Has Many → Participants (rolloutId)

Participants
  └── Has Many → Checklist (participantId)
       └── 18 items per participant

ActivityLog
  └── References → Users (userId)
  └── References → Various targets (targetId)
```

---

## Indexing & Performance

Google Sheets doesn't have traditional database indexes, but the application is optimized by:

1. **Sequential scanning** with early exit when finding by ID
2. **Caching** configuration data in memory
3. **Batch operations** where possible
4. **Column index caching** to avoid repeated header lookups

For optimal performance:
- Keep each sheet under 10,000 rows
- Archive completed rollouts periodically
- Export and archive old activity logs

---

## Backup Recommendations

1. **Manual backups**: File → Make a copy
2. **Version history**: File → Version history → See version history
3. **Export backups**: Use the app's export functions regularly
4. **Automated backups**: Consider using Google Apps Script triggers for scheduled copies

---

## Security Considerations

1. **Password storage**: Passwords are hashed using SHA-256 (not plaintext)
2. **Session tokens**: Expire after 24 hours
3. **Access control**: Enforced at the application level through role checks
4. **Spreadsheet access**: Limit direct spreadsheet access to administrators only
5. **Activity logging**: All significant actions are logged for audit purposes
