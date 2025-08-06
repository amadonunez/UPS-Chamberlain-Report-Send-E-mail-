# UPS Chamberlain Report(Send E-mail)

This is the repository for the Google Apps Script project that automates the process of sending a daily report for the UPS Nogales account.

## Description

This script performs the following actions:

1.  Reads data from a source Google Sheet which contains a QUERY formula to filter records for 'UPS NOGALES'.
2.  Copies this filtered data into a new, temporary Google Sheet.
3.  Exports the temporary sheet as a Microsoft Excel file (.xlsx).
4.  Attaches the Excel file to an email and sends it to a predefined recipient.
5.  Deletes the temporary Google Sheet from Google Drive to avoid clutter.

For complete process documentation, including the source formula and workflow, please refer to:
"IT Documentation: UPS Nogales Automated Report"
https://docs.google.com/document/d/16hT9A1UQLACYSDjdwj_0zhT2ajW8_nsH5RWZtbth8Q8/edit?tab=t.0
