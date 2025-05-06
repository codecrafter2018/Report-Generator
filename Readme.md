# Project Report Generator
# CRM Opportunity Product Report Generator

This application generates Excel reports of opportunity products for users in Dynamics 365 CRM based on their role hierarchy.

## Features

- Generates opportunity product reports for users and their teams
- Processes hierarchical user structures (managers and subordinates)
- Caches lookups for performance optimization
- Uploads reports directly to user records in CRM
- Secure credential handling

## Business Logic Flow

1. *Initialization*:
   - Connects to CRM using secure credentials
   - Starts timer for performance tracking

2. *User Processing*:
   - Fetches filtered users (specific segment, LOB and roles)
   - Starts with users role 
   - For each user role:
     - Processes their opportunity products
     - Processes all subordinates in hierarchy
     - Generates consolidated Excel report
     - Moves up management chain to process managers

3. *Data Collection*:
   - Retrieves shared opportunity products for each user
   - Aggregates geography data from related records
   - Creates unique report lines for each geography combination

4. *Report Generation*:
   - Creates Excel workbook with formatted data
   - Uploads report to user's record in CRM
   - Cleans up temporary files

## Security Requirements

- Requires CRM connection with appropriate privileges
- Never stores credentials in source code
- Uses secure storage for connection strings
- Implements proper error handling

## Configuration

1. Set environment variable:

```bash
# Format:
# CRM_CONNECTION_STRING="Url=[ORG_URL];AuthType=OAuth;Username=[USER];Password=[PASSWORD];ClientId=[CLIENT_ID];RedirectUri=[REDIRECT_URI]"

export CRM_CONNECTION_STRING="Url=YOUR_ORG_URL;AuthType=OAuth;..."
