# Student Academic Report Automation using UiPath

A Robotic Process Automation (RPA) mini project that reads student
records from Microsoft Excel, generates an individual Microsoft Word
academic report for each student, and emails the report to the
corresponding email address.

## Project Overview

Preparing academic reports manually involves repeatedly copying student
details and marks into documents, saving separate files, and sending
them to the correct recipients. This project automates those repetitive
steps using UiPath Studio.

The workflow reads the student data from an Excel workbook, processes
each row, fills a Word report template with the current student's
information, saves the report with a unique filename, and sends it as an
email attachment.

## Features

-   Reads student records from an Excel workbook.
-   Processes records one at a time using **For Each Row in Data
    Table**.
-   Extracts student information into UiPath variables.
-   Copies a Word template for each student.
-   Replaces template placeholders with values from the current Excel
    row.
-   Saves a separate Word report using the student's USN in the
    filename.
-   Sends the generated report to the email address in the same row.

## Technologies Used

-   **UiPath Studio** -- workflow automation.
-   **Microsoft Excel** -- source student records.
-   **Microsoft Word** -- report template and generated reports.
-   **UiPath Mail / SMTP activity** -- email delivery.

## Input Data

The Excel workbook is expected to contain these column headers:

  Column     Purpose
  ---------- ---------------------------------------
  `Name`     Student name
  `USN`      Student registration number
  `year`     Academic year
  `IA1`      Internal Assessment 1 marks
  `IA2`      Internal Assessment 2 marks
  `Total`    Total marks already provided in Excel
  `Branch`   Engineering branch
  `Email`    Recipient email address

> The workflow reads the `Total` value from Excel; it does not calculate
> it.

## Workflow

1.  Open the Excel workbook using **Use Excel File**.
2.  Read the worksheet into a DataTable named `dtStudents`.
3.  Iterate through the records using **For Each Row in Data Table**.
4.  Assign the current row's values to variables:
    -   `studentName`
    -   `studentUSN`
    -   `studentYear`
    -   `ia1`
    -   `ia2`
    -   `total`
    -   `branch`
    -   `email`
5.  Build a unique output path using the student's USN.
6.  Copy the original Word template to the output path.
7.  Open the copied file using **Word Application Scope**.
8.  Replace the placeholders with the current student's values.
9.  After the Word scope finishes, send the generated document as an
    email attachment to the current student's email address.
10. Continue until all rows have been processed.

## Word Template Placeholders

The template should contain the following exact placeholders:

  Placeholder       UiPath variable
  ----------------- -----------------
  `[StudentName]`   `studentName`
  `[USN]`           `studentUSN`
  `[Year]`          `studentYear`
  `[IA1]`           `ia1`
  `[IA2]`           `ia2`
  `[Total]`         `total`
  `[Branch]`        `branch`
  `[Email]`         `email`

The template must remain a clean placeholder-based document. Do not use
a previously generated student report as the source template.

## Output

Each student's report is saved in the `StudentReports` folder with a
filename based on the student's USN, for example:

``` text
StudentReports/
├── <USN>_Report.docx
└── ...
```

The email activity is placed after **Word Application Scope**, but
remains inside the **For Each Row in Data Table** loop. This ensures
that the current student's report is sent to the current student's email
address.

## Project Structure

A recommended repository structure is:

``` text
Student-Academic-Report-Automation/
├── Main.xaml
├── project.json
├── project.uiproj
├── Templates/
│   └── Student_Academic_Report_Template.docx
├── Documentation/
│   ├── Project_Report.docx
│   └── Workflow_Diagram.png
├── Screenshots/
└── README.md
```

Keep private input workbooks, generated reports containing personal
information, local credentials, and machine-specific configuration files
out of the public repository.

## Setup and Execution

1.  Install UiPath Studio on a supported Windows environment.
2.  Clone or download this repository.
3.  Open the UiPath project using its `.uiproj` file or open `Main.xaml`
    in UiPath Studio.
4.  Install or restore the required UiPath activity dependencies used by
    the project (Excel, Word, System and Mail activities).
5.  Prepare an Excel workbook with the column headers listed above. Use
    test data or anonymized data.
6.  Place the workbook and the placeholder-based Word template in the
    paths configured in the workflow.
7.  Ensure the output folder exists, or create it before running the
    workflow.
8.  Configure the mail connection or SMTP credentials securely in
    UiPath. Do not hard-code passwords in the workflow or commit them to
    GitHub.
9.  Run the workflow and verify each generated report and its matching
    email recipient.

> File names and paths in the workflow must match the files in your
> local project. Review the configured Excel path, template path, and
> output path before execution.

## Validation Checklist

-   [ ] Excel Read Range loads all populated student rows.
-   [ ] The DataTable includes the expected column headers.
-   [ ] Each iteration uses values from its current row.
-   [ ] A fresh copy of the original Word template is created for each
    student.
-   [ ] Word Application Scope opens `reportPath`.
-   [ ] All required placeholders are replaced.
-   [ ] The Word document is saved before the email activity runs.
-   [ ] Each report filename uses the correct student's USN.
-   [ ] Each email recipient matches the same Excel row as the attached
    report.

## Security and Privacy

-   Do not commit real student names, USNs, email addresses, marks, or
    generated reports containing personal data.
-   Use anonymized sample records when sharing screenshots or sample
    workbooks.
-   Never commit Gmail passwords, app passwords, API keys, or other
    credentials.
-   Keep private or machine-specific settings out of the repository.

## Future Enhancements

-   Generate PDF versions of the reports.
-   Validate missing or incorrectly formatted student data before
    processing.
-   Add logging for successful and failed report generation and email
    delivery.
-   Add exception handling and retry logic for file and email errors.
-   Generate a final summary of reports created and emails sent.

## Author

**Sanjana R. Bharade**\
Computer Science and Engineering

------------------------------------------------------------------------

*This project was developed as an academic mini project to demonstrate
workflow automation using UiPath.*
