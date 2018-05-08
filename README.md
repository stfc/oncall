# Automating OnCall Spreadsheet
This is a Python script to aid filling in the OnCall spreadsheet. The spreadsheet is used during the 3pm meeting on a Monday to discuss callouts from the past week. 
The script takes data from a TSV file downloaded from RT using a specific search (see 'Using Script') and appends it to the ongoing OnCall spreadsheet (found within Tier 1's section on the TWiki).

## Required Packages
Only openpyxl is needed for this script. The script was written using version 2.5.3 of this module.

## Using Script
There are a couple of steps to using this script as there will be some cleanup to do once the script has run however this makes the process significantly quicker.

### Pre-Script Execution
- Create a custom ticket search in RT:
    - Normal OnCall queue (new or open tickets)
    - Ignore 'NoCall' tickets
    - Ignore 'Tier1_service_Test_Nagios_and_paging_on_host_nagger' tickets (test pager tickets)
    - Ignore 'Downtime Expiry Report' tickets
    - Sort ID by ascending
- Download spreadsheet (.TSV file) from search results
- Put OnCall spreadsheet (found in TWiki within Tier 1 section) and TSV file in same folder as script and then run it

### Post-Script Execution
- For each row added because of the script, double click and press ENTER on the cell in 'Time issued' column. This will execute the formula that deals with the conditional formatting for working hours as this doesn't occur during the script. 
- Fill in any blank cells which the script didn't fill:
    - Host
    - Service
    - People involved
    - Handled by
    - Any helpful comments
- Upload to TWiki
- After the meeting, resolve all tickets from the OnCall queue, except the ones where the situation is still ongoing or haven't been put in the spreadsheet

