# Automating OnCall Spreadsheet
A Python script to aid filling in the OnCall spreadsheet in preparation for the OnCall meeting in RIG using data from a TSV file downloaded from RT using a specific search (see 'Using Script').

## Required Packages
Package | Version
------- | -------
xlrd | 1.1.0
xlutils | 2.0.0

## Using Script
- Create a custom ticket search in RT
    - Normal OnCall queue (new or open tickets)
    - Ignore 'NoCall' tickets
    - Ignore 'Tier1_service_Test_Nagios_and_paging_on_host_nagger' tickets (test pager tickets)
    - Ignore 'Downtime Expiry Report' tickets
- Download spreadsheet (.TSV file) from search results
- Put OnCall spreadsheet (found in TWiki within Tier 1 section) and TSV file in same folder as script then Python script
- Perform usual OnCall meeting protocol regarding the tickets - resolve those that have been cleared up

