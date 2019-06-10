# Automating OnCall Spreadsheet
This is a Python script to aid filling in the OnCall spreadsheet. The spreadsheet is used during the 3pm meeting on a Monday to discuss callouts from the past week.
The script takes data from a TSV file downloaded from RT using a specific search (see 'Using Script') and appends it to the ongoing OnCall spreadsheet (found within Tier 1's section on the TWiki).

## Required Packages
Only openpyxl is needed for this script. The script was written using version 2.5.3 of this module.

## Using Script
There are a couple of steps to using this script as there will be some cleanup to do once the script has executed.

### Pre-Script Execution
- Create a custom ticket search in RT:
    - Queue is OnCall
    - Status is new or open
    - Subject not like 'NoCall'
    - Subject not like 'Test_Nagios' (test tickets)
    - Subject not like 'Downtime Expiry Report'
    - Sort ID by ascending
- Set the first four display columns to the following order:
    - id
    - Subject
    - Status
    - Created
- Download spreadsheet (.TSV file) from search results (press 'Spreadsheet' hyperlink towards top right of the page)
- Put OnCall spreadsheet (found in TWiki within Tier 1 section) and TSV file in same folder as script and then run script

### Post-Script Execution
- Format column D as Time
- Fill in any blank cells the script didn't fill from the following columns:
    - Host
    - Service
    - People involved
    - Handled by
    - Any helpful comments
- Upload to TWiki
- After the meeting, resolve all tickets from the OnCall queue search, except the ones where the situation is still ongoing
