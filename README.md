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
    - Subject not like 'Tier1_service_Test_Nagios_and_paging_on_host_nagger' (test pager tickets)
    - Subject not like 'Downtime Expiry Report'
    - Sort ID by ascending
- Download spreadsheet (.TSV file) from search results (press 'Spreadsheet' hyperlink towards top right of the page)
- Put OnCall spreadsheet (found in TWiki within Tier 1 section) and TSV file in same folder as script and then run script

### Post-Script Execution
- For each row added by the script, double click and press ENTER on the cell in 'Time issued' column. This will execute the formula that deals with the conditional formatting for working hours as this doesn't occur during the script. You will know which rows should be black as these will have the 'Normal Service Restored?'  and the 'Handled by' columns filled in.
- Fill in any blank cells the script didn't fill from the following columns:
    - Host
    - Service
    - People involved
    - Handled by
    - Any helpful comments
- Upload to TWiki
- After the meeting, resolve all tickets from the OnCall queue, except the ones where the situation is still ongoing or haven't been put in the spreadsheet